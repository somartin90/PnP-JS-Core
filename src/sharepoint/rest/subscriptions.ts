"use strict";

import { Queryable, QueryableCollection, QueryableInstance } from "./queryable";

/**
 * Describes a collection of webhook subscriptions
 * 
 */
export class Subscriptions extends QueryableCollection {

    /**
     * Creates a new instance of the Subscriptions class
     * 
     * @param baseUrl - The url or Queryable which forms the parent of this webhook subscriptions collection
     */
    constructor(baseUrl: string | Queryable, path = "subscriptions") {
        super(baseUrl, path);
    }

    /**
     * Returns all the webhook subscriptions or the specified webhook subscription
     *
     */
    public getById(subscriptionId: string): Subscription {
        let subscription = new Subscription(this);
        subscription.concat(`('${subscriptionId}')`);
        return subscription;
    }

    /**
     * Create a new webhook subscription
     *
     */
    public add(notificationUrl: string, expirationDate: string, clientState?: string): Promise<SubscriptionAddResult> {

        let postBody = JSON.stringify({
            "resource": this.toUrl(),
            "notificationUrl": notificationUrl,
            "expirationDateTime": expirationDate,
            "clientState": clientState || "pnp-js-core-subscription",
        });

        return this.post({ body: postBody, headers: { "Content-Type": "application/json" } }).then(result => {

            return { data: result, subscription: this.getById(result.id) };
        });
    }
}

/**
 * Describes a single webhook subscription instance
 * 
 */
export class Subscription extends QueryableInstance {

    /**
     * Creates a new instance of the Subscription class
     * 
     * @param baseUrl - The url or Queryable which forms the parent of this webhook subscription instance
     */
    constructor(baseUrl: string | Queryable, path?: string) {
        super(baseUrl, path);
    }

    /**
     * Update a webhook subscription
     *
     */
    public update(expirationDate: string): Promise<SubscriptionUpdateResult> {

        let postBody = JSON.stringify({
            "expirationDateTime": expirationDate,
        });

        return this.patch({ body: postBody, headers: { "Content-Type": "application/json" } }).then(data => {
            return { data: data, subscription: this };
        });
    }

    /**
     * Remove a webhook subscription
     *
     */
    public delete(): Promise<void> {
        return super.delete();
    }
}

export interface SubscriptionAddResult {
    subscription: Subscription;
    data: any;
}

export interface SubscriptionUpdateResult {
    subscription: Subscription;
    data: any;
}
